from typing import IO

from pptx_editor.content_type.presentationml import PresentationML
from pptx_editor.content_types import ContentTypes
from pptx_editor.parser import Parser
from pptx_editor.part import Part
from pptx_editor.attribute import Attribute

class Presentation(Part):
    content_type = PresentationML.PRESENTATION

    @staticmethod
    def from_zip_file(file: IO):
        parser = Parser(file)
        content_types = ContentTypes.from_file(parser)

class SlideSize(Attribute):
    pass