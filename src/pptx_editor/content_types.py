from zipfile import ZipFile

from lxml.etree import _Element
from lxml import etree

from pptx_editor.part import PartRegistry, Part

class ContentTypes():
    def __init__(self, content_types: bytes):
        self.parts = self.parse_content_types(content_types)

    def parse_content_types(self, content_types_bytes: bytes) -> list[Part]:
        content_types = content_types_bytes
        root_element = etree.fromstring(content_types)
        return self.parse_content_types_tree(root_element)

    def parse_content_types_tree(self, tree: _Element) -> list[Part]:
        parts = []

        for child in tree:
            part = self.parse_content_type(child)
            
            if not part:
                continue

            parts.append(part)
        
        return parts

    def parse_content_type(self, content_type: _Element) -> Part | None:
        q = etree.QName(content_type)

        if q.localname == 'Default':
            part = self.parse_default_content_type(content_type)
        elif q.localname == 'Override':
            part = self.parse_override_content_type(content_type)

        return part

    def parse_default_content_type(self, content_type: _Element):
        pass

    def parse_override_content_type(self, content_type: _Element) -> Part | None:
        content_type_value = content_type.get('ContentType')
        
        if content_type_value is None:
            print("Integrity warning: Override element missing ContentType attribute")
            return None

        part_cls = PartRegistry().get_part_cls(content_type_value)

        if not part_cls:
            print(f"Warning: No registered part class for content type: {content_type_value}")
            return None

        part = part_cls(content_type.get('PartName'))

        return part

    def _parse_content_data(self, zip_file: ZipFile):
        for part in self.parts:
            part._parse_data(zip_file)
