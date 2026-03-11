from typing import IO
from zipfile import ZipFile

from pptx_editor.content_type.presentationml import PresentationML
from pptx_editor.content_types import ContentTypes
from pptx_editor.part import Part
from pptx_editor.attribute import Attribute

class Presentation(Part):
    content_type = PresentationML.PRESENTATION

    @staticmethod
    def from_file(file: IO):
        zip_file = ZipFile(file)
        content_types_file = zip_file.read('[Content_Types].xml')
        content_types = ContentTypes(content_types_file)
        content_types._parse_content_data(zip_file)

class SlideSize(Attribute):
    pass