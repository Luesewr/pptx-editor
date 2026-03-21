from typing import IO

from pptx_editor.content_type.presentationml import PresentationML
from pptx_editor.part import Part
from pptx_editor.attribute import Attribute

class Presentation(Part):
    default_content_type = PresentationML.PRESENTATION
    default_base_path = '/ppt'
    default_part_name = 'presentation.xml'

    @staticmethod
    def from_zip_file(file: IO):
        from pptx_editor.parser import Parser

        parser = Parser(file)
        return parser.parse_zip_file()

    def save_to_buffer(self) -> IO:
        from pptx_editor.writer import Writer

        writer = Writer()
        return writer.write_to_buffer(self)

class SlideSize(Attribute):
    pass