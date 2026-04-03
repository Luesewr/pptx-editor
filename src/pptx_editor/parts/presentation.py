from io import BytesIO
from pathlib import PurePosixPath
from typing import IO
from zipfile import ZipFile

from pptx_editor.content_type.presentationml import PresentationML
from pptx_editor.attribute import Attribute
from pptx_editor.parts.xml_part import XmlPart

class Presentation(XmlPart):
    default_content_type = PresentationML.PRESENTATION
    default_base_path = PurePosixPath('/ppt')
    default_part_name = 'presentation.xml'

    @staticmethod
    def from_zip_file(file: IO):
        from pptx_editor.parser import Parser

        parser = Parser(file)
        return parser.parse_zip_file()

    def save_to_buffer(self) -> IO:
        from pptx_editor.writer import Writer

        buffer = BytesIO()

        with ZipFile(buffer, 'w') as zip_file:
            writer = Writer(zip_file)
            writer.write_to_buffer(self)

        buffer.seek(0)
        return buffer

class SlideSize(Attribute):
    pass