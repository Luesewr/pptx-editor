from io import BytesIO
from pathlib import PurePosixPath
from typing import IO
from zipfile import ZipFile, ZIP_DEFLATED

from pptx_editor.content_type.presentationml import PresentationML
from pptx_editor.attribute import Attribute
from pptx_editor.parts.xml_part import XmlPart

class Presentation(XmlPart):
    default_content_type = PresentationML.PRESENTATION
    default_base_path = PurePosixPath('/ppt')
    default_part_name = 'presentation.xml'

    @staticmethod
    def from_zip_file(file: IO):
        from pptx_editor.parser import _OOXMLParser

        parser = _OOXMLParser(file)
        part = parser.parse_zip_file()

        if not isinstance(part, Presentation):
            raise ValueError('The provided file does not contain a presentation part.')

        return part

    def save_to_buffer(self) -> IO:
        from pptx_editor.writer import _OOXMLWriter

        buffer = BytesIO()

        with ZipFile(buffer, 'w', ZIP_DEFLATED) as zip_file:
            writer = _OOXMLWriter(zip_file)
            writer.write_to_buffer(self)

        buffer.seek(0)
        return buffer

class SlideSize(Attribute):
    pass