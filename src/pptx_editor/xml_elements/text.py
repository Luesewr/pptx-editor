from abc import ABC, abstractmethod
from typing import TypeGuard

from pptx_editor.xml_element import XmlElement
from pptx_editor.exceptions import PowerpointIntegrityError

class TextBody(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    default_name = 'txBody'

    @property
    def paragraphs(self) -> list['Paragraph']:
        attributes = self.get_attributes('p', 'a')

        if not self._is_paragraphs_valid(attributes):
            raise PowerpointIntegrityError('All paragraph attributes must be of type Paragraph.')

        return attributes

    @property
    def paragraph_texts(self) -> str:
        return '\n\n'.join(paragraph.paragraph_text for paragraph in self.paragraphs)

    def _is_paragraphs_valid(self, paragraphs: list[XmlElement]) -> TypeGuard[list['Paragraph']]:
        return all(isinstance(paragraph, Paragraph) for paragraph in paragraphs)


class Paragraph(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_name = 'p'

    @property
    def runs(self) -> list['Run']:
        return [element for element in self.children if isinstance(element, Run)]

    @property
    def paragraph_elements(self) -> list['ParagraphContent']:
        return [element for element in self.children if isinstance(element, ParagraphContent)]

    @property
    def paragraph_text(self) -> str:
        return ''.join(element.content_text for element in self.paragraph_elements)


class ParagraphContent(XmlElement, ABC):
    @property
    @abstractmethod
    def content_text(self) -> str:
        pass

class Run(ParagraphContent):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_name = 'r'

    @property
    def content_text(self) -> str:
        text_attribute = self.get_attribute('t', 'a')

        if text_attribute is None:
            raise PowerpointIntegrityError('The run element is missing the required t element.')

        if not isinstance(text_attribute, Text):
            raise PowerpointIntegrityError('The t element in the run element is not of the expected type.')

        return text_attribute.text

    @content_text.setter
    def content_text(self, value: str) -> None:
        text_attribute = self.get_attribute('t', 'a')

        if text_attribute is None:
            raise PowerpointIntegrityError('The run element is missing the required t element.')

        if not isinstance(text_attribute, Text):
            raise PowerpointIntegrityError('The t element in the run element is not of the expected type.')

        text_attribute.text = value



class Text(XmlElement):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_name = 't'

class Break(ParagraphContent):
    default_namespace = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    default_name = 'br'

    @property
    def content_text(self) -> str:
        return '\n'
