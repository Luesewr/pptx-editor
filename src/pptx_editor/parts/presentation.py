from typing import IO

from pptx_editor.content_type.presentationml import PresentationML
from pptx_editor.part import ReturnPart
from pptx_editor.attribute import Attribute

class Presentation(ReturnPart):
    content_type = PresentationML.PRESENTATION

    @staticmethod
    def from_zip_file(file: IO):
        from pptx_editor.parser import Parser

        parser = Parser(file)
        return parser.parse_zip_file()

class SlideSize(Attribute):
    pass