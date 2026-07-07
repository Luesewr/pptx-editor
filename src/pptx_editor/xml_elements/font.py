from abc import ABC

from pptx_editor.attribute import AttributeStringProperty
from pptx_editor.attributes.text import Typeface
from pptx_editor.xml_element import XmlElement


class AbstractFont(XmlElement, ABC):
    is_abstract = True

    typeface: str = AttributeStringProperty(Typeface, nullable=False)

    def __init_subclass__(cls, **kwargs):
        cls.is_abstract = False
        super().__init_subclass__(**kwargs)

class BulletFont(AbstractFont):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'buFont'

class ComplexScriptFont(AbstractFont):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'cs'

class EastAsianFont(AbstractFont):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'ea'

class LatinFont(AbstractFont):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'latin'

class SymbolFont(AbstractFont):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_prefix = 'a'
    default_name = 'sym'
