from io import BytesIO
from pathlib import PurePosixPath
from typing import IO, TYPE_CHECKING
from zipfile import ZipFile, ZIP_DEFLATED

import pptx_editor.parser
import pptx_editor.writer

from pptx_editor.content_type.presentationml import PresentationML
from pptx_editor.exceptions import PowerpointIntegrityError
from pptx_editor.parts.xml_part import XmlPart
from pptx_editor.parts.slide import Slide
from pptx_editor.xml_elements.id_list import SlideIdList

class Presentation(XmlPart):
    default_content_type = PresentationML.PRESENTATION
    default_base_path = PurePosixPath('/ppt')
    default_part_name = 'presentation'

    @property
    def slides(self) -> list['Slide']:
        return self._slide_id_list.slides

    @property
    def _slide_id_list(self) -> 'SlideIdList':
        slide_id_list = self.get_element_by_type(SlideIdList)

        if slide_id_list is None:
            raise PowerpointIntegrityError('The presentation part is missing the required sldIdLst element.')

        if not isinstance(slide_id_list, SlideIdList):
            raise PowerpointIntegrityError('The sldIdLst element in the presentation part is not of the expected type.')

        return slide_id_list

    def add_slide(self, slide: 'Slide', index: int | None = None) -> None:
        self._slide_id_list.add_slide(slide, index)

    @staticmethod
    def from_zip_file(file: IO):
        parser = pptx_editor.parser._OOXMLParser(file)
        part = parser.parse_zip_file()

        if not isinstance(part, Presentation):
            raise ValueError('The provided file does not contain a presentation part.')

        return part

    def save_to_buffer(self) -> IO:
        buffer = BytesIO()

        with ZipFile(buffer, 'w', ZIP_DEFLATED) as zip_file:
            writer = pptx_editor.writer._OOXMLWriter(zip_file)
            writer.write_to_buffer(self)

        buffer.seek(0)
        return buffer

    def save_to_file(self, file: IO) -> None:
        with ZipFile(file, 'w', ZIP_DEFLATED) as zip_file:
            writer = pptx_editor.writer._OOXMLWriter(zip_file)
            writer.write_to_buffer(self)
