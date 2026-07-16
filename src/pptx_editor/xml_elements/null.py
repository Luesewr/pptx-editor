from io import BytesIO
from typing import Self, TypeVar, Generic, TypeGuard

from pptx_editor.modes.xml_element import AddMode
from pptx_editor.writer import _OOXMLWriter
from pptx_editor.xml_element import XmlElement

T = TypeVar('T', bound=XmlElement)

class NullElement(XmlElement, Generic[T]):
    default_namespace = None
    default_prefix = None
    default_name = None

    def __init__(self, element_type: type[T], parent: XmlElement | None = None):
        super().__init__()
        self.element_type = element_type
        self.parent = parent

    def create_if_null(self, element_type: type[T] | None = None, add_mode: AddMode = AddMode.SORT) -> T:
        if element_type is None:
            element_type = self.element_type

        element = element_type(part=self.parent.part)

        if self.parent is not None:
            self.parent.auto_add_element(element, add_mode=add_mode)

        return element

    def _to_xml(self, writer: '_OOXMLWriter', buffer: BytesIO) -> str:
        pass
