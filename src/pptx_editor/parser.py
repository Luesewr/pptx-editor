from io import BytesIO
from pathlib import PurePosixPath
from zipfile import ZipFile
from typing import IO
from collections import defaultdict

from lxml import etree

from pptx_editor.attribute import Attribute
from pptx_editor.content_types import ContentTypes
from pptx_editor.exceptions import PowerpointIntegrityError
from pptx_editor.part import Part, PartRegistry
from pptx_editor.parts import package
from pptx_editor.registries.attribute import AttributeRegistry
from pptx_editor.xml_element import XmlElement, XmlElementRegistry

class _OOXMLParser:
    attribute_registry = AttributeRegistry()
    element_registry = XmlElementRegistry()
    part_registry = PartRegistry()

    def __init__(self, file: IO, return_location: str = '/ppt/presentation.xml'):
        self.zip_file = ZipFile(file)
        self.parts: dict[PurePosixPath, 'Part'] = {}
        self.content_types: ContentTypes | None = None
        self.return_location = PurePosixPath(return_location)

        self.package: 'package.Package' = package.Package(None)
        self.main_part: 'Part' | None = None


    def parse_zip_file(self) -> 'Part':
        if str(self.return_location).lstrip('/') not in self.zip_file.namelist():
            raise PowerpointIntegrityError(f"Integrity warning: Main presentation part {self.return_location} not found in zip file")

        self.content_types = ContentTypes._from_file(self)

        self.package._parse_relationships(self, PurePosixPath('/'))

        return_part = self.get_part(self.return_location)

        if return_part is None:
            raise PowerpointIntegrityError(f"Integrity warning: Main presentation part {self.return_location} not found")

        return return_part

    def parse_part_from_file(self, file_path: PurePosixPath):
        content_type, is_default = self.get_content_type(file_path)

        part_cls = _OOXMLParser.part_registry.get_part_cls(content_type)
        part = part_cls._from_file(self, file_path, content_type, is_default)

        return part

    @staticmethod
    def parse_part_from_xml(part: 'Part'):
        if part._data is None:
            return None, None

        context = etree.iterparse(
            BytesIO(part._data),
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
                namespace, prefix, name = _OOXMLParser._process_tag(value.tag, ns_map)

                element = path_stack.pop()
                declared_namespaces = ns_stack.pop()
                children = tuple(child_stack.pop())

                parsed_attributes = tuple(
                    _OOXMLParser.parse_attribute_from_item(part, element.nsmap, name, str(key), str(value))
                    for key, value
                    in element.attrib.items()
                )

                parsed_element = _OOXMLParser.parse_element_from_xml(part, element, parsed_attributes, children, declared_namespaces)

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
    def parse_element_from_xml(part: 'Part', xml: etree._Element, attributes: tuple['Attribute'], children: tuple['XmlElement'], ns_declarations: dict[str, dict[str | None, str]] | None) -> 'XmlElement':
        q = etree.QName(xml)
        namespace = q.namespace or None
        name = q.localname
        prefix = xml.prefix or None
        text = xml.text or None
        tail = xml.tail or None

        element_cls = _OOXMLParser.element_registry.get_element_cls(namespace, name)
        element = element_cls(name, prefix, attributes, children, text, tail, part, ns_declarations)
        return element

    @staticmethod
    def parse_attribute_from_item(part: 'Part', namespaces: dict[str | None, str], element_name: str, name: str, value: str):
        q = etree.QName(name)
        namespace = q.namespace if q.namespace else None
        attribute_cls = _OOXMLParser.attribute_registry.get_attribute_value_cls(namespace, name, element_name)
        attribute_value = attribute_cls._from_item(part, namespaces, name, value)

        return attribute_value

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

    def get_content_type(self, file_path: PurePosixPath) -> tuple[str, bool]:
        if self.content_types is None:
            raise PowerpointIntegrityError("Content types not loaded")

        override_content_type = self.content_types.get_override_content_type(file_path)

        if override_content_type is not None:
            return override_content_type, False

        default_content_type = self.content_types.get_default_content_type(file_path)

        if default_content_type is not None:
            return default_content_type, True

        raise PowerpointIntegrityError(f"Integrity warning: No content type found for file {file_path}")

    def read_file(self, file_path: PurePosixPath) -> bytes:
        if file_path.is_absolute():
            file_path = file_path.relative_to(file_path.anchor)

        return self.zip_file.read(file_path.as_posix())

    def add_part(self, file_path: PurePosixPath | None, part: 'Part'):
        if file_path:
            self.parts[file_path] = part

    def has_part(self, file_path: PurePosixPath) -> bool:
        return file_path in self.parts

    def get_part(self, file_path: PurePosixPath):
        return self.parts.get(file_path)
