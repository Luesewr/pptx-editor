from pptx_editor.parts.slide import Slide
from pptx_editor.properties.id_list import PartIdListProperty
from pptx_editor.xml_element import XmlElement

class AbstractIdList(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = None
    id_class = None
    relationship_type = None
    is_abstract = True

    def __init_subclass__(cls, **kwargs):
        cls.is_abstract = False
        super().__init_subclass__(**kwargs)

        missing_id_class = cls.id_class is None
        missing_relationship_type = cls.relationship_type is None

        if missing_id_class:
            raise TypeError(f"Subclasses of AbstractIdList must define an 'id_class' attribute. {cls.__name__} does not.")

        if missing_relationship_type:
            raise TypeError(f"Subclasses of AbstractIdList must define a 'relationship_type' attribute. {cls.__name__} does not.")


class AbstractId(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = None
    is_abstract = True

    def __init_subclass__(cls, **kwargs):
        cls.is_abstract = False
        super().__init_subclass__(**kwargs)


class SlideId(AbstractId):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = 'sldId'

class SlideIdList(AbstractIdList):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_prefix = 'p'
    default_name = 'sldIdLst'
    id_class = SlideId
    relationship_type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'

    slides = PartIdListProperty(Slide)

    @property
    def _slide_ids(self) -> list['SlideId']:
        return self.get_elements_by_type(SlideId)
