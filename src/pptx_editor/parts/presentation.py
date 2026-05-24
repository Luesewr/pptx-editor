from io import BytesIO
from pathlib import PurePosixPath
from typing import IO, TYPE_CHECKING
from zipfile import ZipFile, ZIP_DEFLATED

from pptx_editor.content_type.presentationml import PresentationML
from pptx_editor.exceptions import PowerpointIntegrityError
from pptx_editor.parts.xml_part import XmlPart
from pptx_editor.parts.slide import Slide

if TYPE_CHECKING:
    from pptx_editor.attributes.id_list import SlideIdList

class Presentation(XmlPart):
    default_content_type = PresentationML.PRESENTATION
    default_base_path = PurePosixPath('/ppt')
    default_part_name = 'presentation.xml'

    @property
    def slides(self) -> list['Slide']:
        return self._slide_id_list.slides()

    @property
    def _slide_id_list(self) -> 'SlideIdList':
        from pptx_editor.attributes.id_list import SlideIdList
        slide_id_list = self.get_attribute('sldIdLst', 'p')

        if slide_id_list is None:
            raise PowerpointIntegrityError('The presentation part is missing the required sldIdLst element.')

        if slide_id_list is not None and not isinstance(slide_id_list, SlideIdList):
            raise PowerpointIntegrityError('The sldIdLst element in the presentation part is not of the expected type.')

        return slide_id_list

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
