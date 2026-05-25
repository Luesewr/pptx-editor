from typing import TYPE_CHECKING, TypeGuard

from pptx_editor.xml_element import XmlElement
from pptx_editor.exceptions import PowerpointIntegrityError

if TYPE_CHECKING:
    from pptx_editor.parts.slide import Slide

class SlideId(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_name = 'sldId'

    def get_slide(self) -> 'Slide':
        from pptx_editor.relationship import Relationship
        from pptx_editor.attributes.relation_attribute import RelationshipValue
        from pptx_editor.parts.slide import Slide

        slide_relationship_value = self.get_value('id', 'r')

        if isinstance(slide_relationship_value, RelationshipValue) and isinstance(slide_relationship_value.value, Relationship) and isinstance(slide_relationship_value.value.target, Slide):
            return slide_relationship_value.value.target

        raise PowerpointIntegrityError('The sldId element does not have a valid relationship to a slide.')

class SlideIdList(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_name = 'sldIdLst'

    def slides(self) -> list['Slide']:
        return [slide_id.get_slide() for slide_id in self._slide_ids()]

    def _slide_ids(self) -> list[SlideId]:
        attributes = self.get_attributes('sldId', 'p')

        if not self._is_slide_ids_valid(attributes):
            raise PowerpointIntegrityError('All sldId attributes must be of type SlideId.')

        return attributes

    def _is_slide_ids_valid(self, slide_ids: list[XmlElement]) -> TypeGuard[list[SlideId]]:
        return all(isinstance(slide_id, SlideId) for slide_id in slide_ids)
