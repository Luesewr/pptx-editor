import random

from pptx_editor.attribute import Attribute
from pptx_editor.attributes.relation_attribute import RelationshipAttribute
from pptx_editor.exceptions import PowerpointIntegrityError
from pptx_editor.modes.xml_element import AddMode
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

    def add_slide(self, slide: 'Slide', index: int | None = None, add_mode: AddMode = AddMode.SORT) -> None:
        if index is None:
            index = len(self._slide_ids)

        if index < 0 or index > len(self._slide_ids):
            raise IndexError("Index out of range.")

        relationship = Relationship('http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide', slide, self.part)
        self.part.add_relationship(relationship)

        slide_relationship = RelationshipAttribute(relationship, name='id')
        unique_slide_id = Attribute(str(random.randint(256, 2147483647)), name='id')
        slide_id = SlideId(attributes=(unique_slide_id, slide_relationship,))

        self.auto_add_element(slide_id, index=index, add_mode=add_mode)

    @property
    def slides(self) -> list['Slide']:
        return [slide_id.get_slide() for slide_id in self._slide_ids]

    @property
    def _slide_ids(self) -> list['SlideId']:
        return self.get_elements_by_type(SlideId)
