from pathlib import PurePosixPath
from typing import TYPE_CHECKING

from pptx_editor.content_type.presentationml import PresentationML
from pptx_editor.exceptions import PowerpointIntegrityError
from pptx_editor.parts.xml_part import XmlPart

if TYPE_CHECKING:
    from pptx_editor.attributes.slide_data import CommonSlideData
    from pptx_editor.attributes.shape import Shape

class Slide(XmlPart):
    default_content_type = PresentationML.SLIDE
    default_base_path = PurePosixPath('/ppt/slides')
    default_part_name = 'slide{i}.xml'

    @property
    def shapes(self) -> list['Shape']:
        return self._common_slide_data.shapes()

    @property
    def _common_slide_data(self) -> 'CommonSlideData':
        from pptx_editor.attributes.slide_data import CommonSlideData
        slide_id_list = self.get_attribute('cSld', 'p')

        if slide_id_list is None:
            raise PowerpointIntegrityError('The presentation part is missing the required cSld element.')

        if slide_id_list is not None and not isinstance(slide_id_list, CommonSlideData):
            raise PowerpointIntegrityError('The cSld element in the presentation part is not of the expected type.')

        return slide_id_list
