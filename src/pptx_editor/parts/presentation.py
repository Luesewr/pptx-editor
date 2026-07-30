from io import BytesIO
from pathlib import PurePosixPath
from typing import IO
from zipfile import ZipFile, ZIP_DEFLATED

import pptx_editor.parser
import pptx_editor.writer

from pptx_editor.content_type.presentationml import PresentationML
from pptx_editor.parts.masters import NotesMaster, SlideMaster
from pptx_editor.parts.slide import Slide
from pptx_editor.parts.xml_part import XmlPart
from pptx_editor.properties.xml_element import RequiredXmlElementProperty
from pptx_editor.xml_elements.presentation import SlideMasterIdList, NotesMasterIdList, SlideIdList


class Presentation(XmlPart):
    default_content_type = PresentationML.PRESENTATION
    default_base_path = PurePosixPath('/ppt')
    default_part_name = 'presentation'

    _slide_master_id_list = RequiredXmlElementProperty(SlideMasterIdList)
    _notes_master_id_list = RequiredXmlElementProperty(NotesMasterIdList)
    _slide_id_list = RequiredXmlElementProperty(SlideIdList)

    @property
    def slides(self) -> list['Slide']:
        return self._slide_id_list.slides

    @slides.setter
    def slides(self, value: list['Slide']) -> None:
        self._slide_id_list.slides = value

    @property
    def slide_masters(self) -> list['SlideMaster']:
        return self._slide_master_id_list.slide_masters

    @slide_masters.setter
    def slide_masters(self, value: list['SlideMaster']) -> None:
        self._slide_master_id_list.slide_masters = value

    @property
    def notes_masters(self) -> list['NotesMaster']:
        return self._notes_master_id_list.notes_masters

    @notes_masters.setter
    def notes_masters(self, value: list['NotesMaster']) -> None:
        self._notes_master_id_list.notes_masters = value

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
