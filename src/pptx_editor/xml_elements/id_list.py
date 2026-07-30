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
