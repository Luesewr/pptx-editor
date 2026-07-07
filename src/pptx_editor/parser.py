from collections import defaultdict
from pathlib import PurePosixPath
from zipfile import ZipFile
from typing import IO

from lxml import etree

import pptx_editor.parts.package as package_part
from pptx_editor.attribute import AttributeRegistry
from pptx_editor.content_types import ContentTypes
from pptx_editor.exceptions import PowerpointIntegrityError
from pptx_editor.part import Part, PartRegistry
from pptx_editor.relationship import Relationship
from pptx_editor.xml_element import XmlElement, XmlElementRegistry

class _OOXMLParser:
    def __init__(self, file: IO, return_location: str = '/ppt/presentation.xml'):
        self.zip_file = ZipFile(file)
        self.parts: dict[PurePosixPath, 'Part'] = {}
        self.relationships: dict[PurePosixPath, dict[str, Relationship]] = defaultdict(dict)
        self.content_types: ContentTypes | None = None
        self.return_location = PurePosixPath(return_location)

        self.package: 'package_part.Package' = package_part.Package(None)
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

        part_cls = PartRegistry().get_part_cls(content_type)
        part = part_cls._from_file(self, file_path, content_type, is_default)

        return part

    def parse_element_from_xml(self, file_path: PurePosixPath | None, xml: etree._Element, ns_declarations: dict[str | None, str]) -> 'XmlElement':
        registry = XmlElementRegistry()
        q = etree.QName(xml)
        namespace = q.namespace if q.namespace else None
        name = q.localname

        element_cls = registry.get_element_cls(namespace, name)
        element = element_cls._from_xml(self, file_path, xml, ns_declarations)
        return element

    def parse_attribute_from_item(self, file_path: PurePosixPath | None, namespaces: dict[str | None, str], element_name: str, name: str, value: str):
        registry = AttributeRegistry()
        q = etree.QName(name)
        namespace = q.namespace if q.namespace else None
        attribute_cls = registry.get_attribute_value_cls(namespace, name, element_name)
        attribute_value = attribute_cls._from_item(self, file_path, namespaces, name, value)

        return attribute_value

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

    def add_relationship(self, file_path: PurePosixPath | None, relationship_id: str | None, relationship: 'Relationship'):
        if file_path and relationship_id:
            self.relationships[file_path][relationship_id] = relationship

    def has_relationship(self, file_path: PurePosixPath, relationship_id: str) -> bool:
        return relationship_id in self.relationships[file_path]

    def get_part(self, file_path: PurePosixPath):
        return self.parts.get(file_path)

    def get_relationship(self, file_path: PurePosixPath, relationship_id: str) -> 'Relationship | None':
        return self.relationships[file_path].get(relationship_id)
