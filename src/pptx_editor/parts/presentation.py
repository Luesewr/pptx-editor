from typing import IO, TYPE_CHECKING

from pptx_editor.content_type.presentationml import PresentationML
from pptx_editor.part import Part
from pptx_editor.attribute import Attribute

if TYPE_CHECKING:
    from pptx_editor.relationship import Relationship

class Presentation(Part):
    content_type = PresentationML.PRESENTATION

    def __init__(self, file_path: str, content_type: str):
        super().__init__(file_path, content_type)
        self.main_relationships: list['Relationship'] | None = None

    @staticmethod
    def from_zip_file(file: IO):
        from pptx_editor.parser import Parser

        parser = Parser(file)
        return parser.parse_zip_file()

class SlideSize(Attribute):
    pass