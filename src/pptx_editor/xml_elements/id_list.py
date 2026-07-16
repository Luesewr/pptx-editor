from pptx_editor.attributes.relation_attribute import RelationshipAttribute
from pptx_editor.exceptions import PowerpointIntegrityError
from pptx_editor.parts.slide import Slide
from pptx_editor.relationship import Relationship
from pptx_editor.xml_element import XmlElement

class SlideId(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = 'sldId'

    def get_slide(self) -> 'Slide':
        slide_relationship_value = self.get_attribute_by_type(RelationshipAttribute)

        if isinstance(slide_relationship_value, RelationshipAttribute) and isinstance(slide_relationship_value.value, Relationship) and isinstance(slide_relationship_value.value.target, Slide):
            return slide_relationship_value.value.target

        raise PowerpointIntegrityError('The sldId element does not have a valid relationship to a slide.')

class SlideIdList(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = 'sldIdLst'

    @property
    def slides(self) -> list['Slide']:
        return [slide_id.get_slide() for slide_id in self._slide_ids]

    @property
    def _slide_ids(self) -> list['SlideId']:
        return self.get_elements_by_type(SlideId)
